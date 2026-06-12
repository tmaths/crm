"""
PDF Report Generator with Template Styling Support

This module generates PDF reports using ReportLab with styling from template_structure.json.
Much faster than Excel with cell merging (~5-10 minutes vs 7+ hours for 300k records).
"""

import json
import os
from decimal import Decimal
from datetime import datetime
from io import BytesIO

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.platypus.flowables import KeepTogether
from reportlab.pdfgen import canvas
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from django.conf import settings


class TemplateStyleParser:
    """Parse template_structure.json and extract styling information."""
    
    def __init__(self, json_path=None):
        if json_path is None:
            json_path = os.path.join(settings.BASE_DIR.parent, 'template_structure.json')
        
        with open(json_path, 'r') as f:
            self.template_data = json.load(f)
        
        self.cells = self.template_data.get('cells', [])
        self.merged_cells = self.template_data.get('merged_cells', [])
        
    def get_cell_style(self, row, col):
        """Get styling for a specific cell."""
        for cell in self.cells:
            if cell['row'] == row and cell['col'] == col:
                return cell
        return None
    
    def parse_color(self, color_str):
        """Convert Excel color format (FFFFFFFF) to ReportLab color."""
        if not color_str or color_str == "00000000":
            return None
        
        # Remove FF prefix if present (alpha channel)
        if color_str.startswith('FF'):
            color_str = color_str[2:]
        
        try:
            r = int(color_str[0:2], 16) / 255.0
            g = int(color_str[2:4], 16) / 255.0
            b = int(color_str[4:6], 16) / 255.0
            return colors.Color(r, g, b)
        except:
            return None
    
    def get_alignment(self, alignment_dict):
        """Convert Excel alignment to ReportLab alignment."""
        if not alignment_dict:
            return TA_CENTER
        
        h_align = alignment_dict.get('horizontal', 'center')
        if h_align == 'left':
            return TA_LEFT
        elif h_align == 'right':
            return TA_RIGHT
        else:
            return TA_CENTER


class PDFReportGenerator:
    """Generate styled PDF reports from data."""
    
    def __init__(self, template_json_path=None):
        self.style_parser = TemplateStyleParser(template_json_path)
        self.styles = getSampleStyleSheet()
        
        # Create custom styles based on template
        self._create_custom_styles()
    
    def _create_custom_styles(self):
        """Create paragraph styles from template JSON."""
        # Header style
        self.styles.add(ParagraphStyle(
            name='ReportHeader',
            parent=self.styles['Heading1'],
            fontSize=12,
            textColor=colors.HexColor('#504C65'),
            alignment=TA_RIGHT,
            fontName='Helvetica-Bold'
        ))
        
        # Billing header
        self.styles.add(ParagraphStyle(
            name='BillingHeader',
            parent=self.styles['Heading2'],
            fontSize=12,
            textColor=colors.HexColor('#F0720C'),
            alignment=TA_LEFT,
            fontName='Helvetica-Bold'
        ))
        
        # Table cell style
        self.styles.add(ParagraphStyle(
            name='TableCell',
            parent=self.styles['Normal'],
            fontSize=6.5,
            textColor=colors.HexColor('#504C65'),
            alignment=TA_LEFT,
            fontName='Helvetica',
            leading=8,  # Line height for better readability
            wordWrap='CJK'  # Enable word wrapping
        ))
        
        # Small text
        self.styles.add(ParagraphStyle(
            name='SmallText',
            parent=self.styles['Normal'],
            fontSize=6,
            textColor=colors.HexColor('#504C65'),
            alignment=TA_RIGHT,
            fontName='Helvetica-Bold'
        ))
    
    def generate_report(self, subscriber_name, start_date_display, end_date_display, 
                       summary_bills=None, product_data=None, username="", output_path=None):
        """
        Generate a PDF report with styling from template.
        
        Args:
            subscriber_name: Name of the subscriber
            start_date_display: Start date as string (DD/MM/YYYY)
            end_date_display: End date as string (DD/MM/YYYY)
            summary_bills: Dict of billing summary
            product_data: List of product records
            username: User who generated the report
            output_path: Path to save PDF (if None, returns BytesIO)
        
        Returns:
            Path to saved PDF or BytesIO object
        """
        # Create PDF buffer or file
        if output_path:
            pdf_file = output_path
        else:
            pdf_file = BytesIO()
        
        # Create document with landscape orientation for wide tables
        doc = SimpleDocTemplate(
            pdf_file,
            pagesize=landscape(A4),
            rightMargin=20*mm,
            leftMargin=20*mm,
            topMargin=20*mm,
            bottomMargin=20*mm
        )
        
        # Build content
        story = []
        
        # Header
        story.append(Paragraph(
            f"FirstCentral NIGERIA - BILLING DETAILS - {subscriber_name}",
            self.styles['ReportHeader']
        ))
        story.append(Spacer(1, 5*mm))
        
        # Contact info
        contact_info = "No 14 Sumbo Jibowu Street, Ikoyi, Lagos, Nigeria<br/>" \
                      "Tel No:+234 (1) 453 4908, +234 (909) 114 1981<br/>" \
                      "website:www.firstcentralcreditbureau.com<br/>" \
                      "Email: info@firstcentralcreditbureau.com"
        story.append(Paragraph(contact_info, self.styles['SmallText']))
        story.append(Spacer(1, 5*mm))
        
        # Date range
        date_range = f"REPORT GENERATED FOR RECORDS BETWEEN {start_date_display} and {end_date_display}"
        story.append(Paragraph(date_range, self.styles['Normal']))
        story.append(Spacer(1, 10*mm))
        
        # Billing section
        if summary_bills:
            story.extend(self._create_billing_section(summary_bills))
            story.append(Spacer(1, 10*mm))
        
        # Product details section
        if product_data:
            story.extend(self._create_product_section(product_data))
        
        # Footer
        story.append(Spacer(1, 10*mm))
        footer_text = f"Report Generated by: {username}"
        footer_style = ParagraphStyle(
            'Footer',
            parent=self.styles['Normal'],
            fontSize=10,
            textColor=colors.HexColor('#7F7F7F'),
            alignment=TA_CENTER,
            fontName='Helvetica-BoldOblique'
        )
        story.append(Paragraph(footer_text, footer_style))
        
        # Build PDF
        doc.build(story)
        
        if output_path:
            return output_path
        else:
            pdf_file.seek(0)
            return pdf_file
    
    def _create_billing_section(self, summary_bills):
        """Create the billing summary section."""
        elements = []
        
        # Section header
        elements.append(Paragraph("Summary Bills", self.styles['BillingHeader']))
        elements.append(Spacer(1, 3*mm))
        
        # Billing table data
        billing_data = [
            ['Product', 'Quantity', 'Rate', 'Amount'],
        ]
        
        # Product mapping
        products = {
            'consumer_snap_check': 'Consumer Snap Check',
            'consumer_basic_trace': 'Consumer Basic Trace',
            'consumer_basic_credit': 'Consumer Basic Credit',
            'consumer_detailed_credit': 'Consumer Detailed Credit',
            'xscore_consumer_detailed_credit': 'X-Score Consumer Detailed Credit',
            'commercial_basic_trace': 'Commercial Basic Trace',
            'commercial_detailed_credit': 'Commercial Detailed Credit',
            'enquiry_report': 'Enquiry Report',
            'consumer_dud_cheque': 'Consumer Dud Cheque',
            'commercial_dud_cheque': 'Commercial Dud Cheque',
            'director_basic_report': 'Director Basic Report',
            'director_detailed_report': 'Director Detailed Report',
        }
        
        total_amount = Decimal('0.00')
        for key, name in products.items():
            qty = summary_bills.get(key, 0)
            rate = summary_bills.get(f'{key}_rate', Decimal('0.00'))
            amount = Decimal(qty) * rate
            total_amount += amount
            
            billing_data.append([
                name,
                str(qty),
                f"₦{rate:,.2f}",
                f"₦{amount:,.2f}"
            ])
        
        # Totals
        vat_amount = total_amount * Decimal('0.075')
        amount_due = total_amount + vat_amount
        
        billing_data.append(['', '', 'Total:', f"₦{total_amount:,.2f}"])
        billing_data.append(['', '', 'VAT (7.5%):', f"₦{vat_amount:,.2f}"])
        billing_data.append(['', '', 'Amount Due:', f"₦{amount_due:,.2f}"])
        
        # Create table
        table = Table(billing_data, colWidths=[120*mm, 30*mm, 40*mm, 40*mm])
        
        # Apply styling
        table.setStyle(TableStyle([
            # Header row
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F5F5F5')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor('#504C65')),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8.25),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            
            # Data rows
            ('FONTNAME', (0, 1), (-1, -4), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 1), (-1, -4), 8.25),
            ('TEXTCOLOR', (0, 1), (-1, -4), colors.HexColor('#504C65')),
            ('ALIGN', (0, 1), (0, -4), 'LEFT'),
            ('ALIGN', (1, 1), (-1, -4), 'RIGHT'),
            
            # Total rows
            ('FONTNAME', (0, -3), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, -3), (-1, -1), 9),
            ('ALIGN', (0, -3), (-1, -1), 'RIGHT'),
            
            # Borders
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#504C65')),
            ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#504C65')),
        ]))
        
        elements.append(table)
        return elements
    
    def _create_product_section(self, product_data):
        """Create product details section with pagination for large datasets."""
        elements = []
        
        # Group by product name
        grouped_products = {}
        for record in product_data:
            product_name = record.get('ProductName', 'Unknown')
            if product_name not in grouped_products:
                grouped_products[product_name] = []
            grouped_products[product_name].append(record)
        
        # Process each product group
        for product_idx, (product_name, records) in enumerate(sorted(grouped_products.items())):
            if product_idx > 0:
                elements.append(PageBreak())
            
            # Product header
            elements.append(Paragraph(f"Product: {product_name}", self.styles['BillingHeader']))
            elements.append(Spacer(1, 5*mm))
            
            # Product table header with ALL columns and full names
            header_row = ['Serial\nNumber', 'Branch ID', 'Unique Tracking\nNumber', 'Subscriber\nName', 
                         'System User', 'Subscriber\nEnquiry Date', 'Product', 'Details Viewed\nDate', 'Search Output']
            
            table_data = []
            ROWS_PER_PAGE = 45  # Create new table every 45 rows for better pagination
            
            # Add ALL records (no limit)
            for idx, record in enumerate(records, 1):  # Process ALL records
                enq_date = record.get('SubscriberEnquiryDate', '')
                if enq_date and not isinstance(enq_date, str):
                    enq_date = enq_date.strftime('%Y-%m-%d')
                
                view_date = record.get('DetailsViewedDate', '')
                if view_date and not isinstance(view_date, str):
                    view_date = view_date.strftime('%Y-%m-%d')
                
                # Get full values without truncation
                # BranchID and UniqueTrackingNumber are placeholders from template (not in DB)
                branch_id = ''  # Placeholder column
                unique_tracking = ''  # Placeholder column
                subscriber_name = str(record.get('SubscriberName', ''))
                system_user = str(record.get('SystemUser', ''))
                product_name_cell = str(record.get('ProductName', ''))
                search_output = str(record.get('SearchOutput', ''))
                
                # Wrap text in Paragraphs for proper wrapping
                table_data.append([
                    Paragraph(str(idx), self.styles['TableCell']),
                    Paragraph(branch_id, self.styles['TableCell']),
                    Paragraph(unique_tracking, self.styles['TableCell']),
                    Paragraph(subscriber_name, self.styles['TableCell']),
                    Paragraph(system_user, self.styles['TableCell']),
                    Paragraph(str(enq_date), self.styles['TableCell']),
                    Paragraph(product_name_cell, self.styles['TableCell']),
                    Paragraph(str(view_date), self.styles['TableCell']),
                    Paragraph(search_output, self.styles['TableCell'])
                ])
                
                # Create table every ROWS_PER_PAGE for memory efficiency and proper pagination
                if len(table_data) >= ROWS_PER_PAGE:
                    # Insert header at the beginning
                    table_data.insert(0, header_row)
                    table = self._create_product_table(table_data)
                    elements.append(table)
                    elements.append(Spacer(1, 3*mm))
                    # Reset table data
                    table_data = []
            
            # Add remaining rows
            if table_data:
                # Insert header at the beginning
                table_data.insert(0, header_row)
                table = self._create_product_table(table_data)
                elements.append(table)
            
            # Show total count
            total_note = f"Total Records: {len(records):,}"
            elements.append(Spacer(1, 3*mm))
            note_style = ParagraphStyle(
                'TotalNote',
                parent=self.styles['Normal'],
                fontSize=9,
                textColor=colors.HexColor('#504C65'),
                alignment=TA_RIGHT,
                fontName='Helvetica-Bold'
            )
            elements.append(Paragraph(total_note, note_style))
        
        return elements
    
    def _create_product_table(self, table_data):
        """Create a styled product table with proper column widths and text wrapping."""
        # Adjusted column widths for landscape A4 (297mm width, minus margins ~257mm available)
        # Serial: 12mm, Branch: 18mm, Tracking: 22mm, Subscriber: 35mm, User: 30mm, 
        # Enq Date: 22mm, Product: 35mm, View Date: 22mm, Output: 61mm
        table = Table(table_data, colWidths=[12*mm, 18*mm, 22*mm, 35*mm, 30*mm, 22*mm, 35*mm, 22*mm, 61*mm])
        
        table.setStyle(TableStyle([
            # Header
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F5F5F5')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor('#504C65')),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 7),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
            
            # Data rows
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 6.5),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#504C65')),
            ('ALIGN', (0, 1), (0, -1), 'CENTER'),  # Serial Number centered
            ('ALIGN', (1, 1), (-1, -1), 'LEFT'),   # All other columns left-aligned
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),   # Vertical align to top for wrapped text
            
            # Borders
            ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.HexColor('#504C65')),
            
            # Row height - allow auto-sizing for wrapped text
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#FAFAFA')]),
        ]))
        
        return table
