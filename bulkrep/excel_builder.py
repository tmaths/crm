"""
Excel report builder using xlsxwriter for high-performance report generation.

Key advantages over openpyxl:
- Format objects are created once and reused by reference (no per-cell overhead)
- Writes stream directly to disk (no in-memory cell model)
- merge_range() writes value + format in one call
- 3-5x faster for large datasets
"""

import os
import json
import logging
import tempfile
from decimal import Decimal
from functools import lru_cache

import xlsxwriter
from django.conf import settings

logger = logging.getLogger(__name__)

# Excel hard limit: 1,048,576 rows per sheet. We stop at 1,000,000 to leave
# breathing room for headers on the overflow sheet.
MAX_ROWS_PER_SHEET = 1_000_000

# Template structure cache
_template_cache = None

def load_template_structure():
    """Load and cache the template structure from JSON."""
    global _template_cache
    if _template_cache is None:
        template_path = os.path.join(settings.BASE_DIR.parent, 'template_structure.json')
        with open(template_path, 'r', encoding='utf-8') as f:
            _template_cache = json.load(f)
    return _template_cache


def prepare_dynamic_values(subscriber_name, start_date_display, end_date_display):
    """Prepare value overrides for dynamic content in the template."""
    return {
        'H2': f"FirstCentral NIGERIA - BILLING DETAILS - {subscriber_name}",
        'B6': f"Report Generated for Records Between {start_date_display} and {end_date_display}",
    }


def _convert_color(color_str):
    """Convert AARRGGBB color string to #RRGGBB for xlsxwriter."""
    if not color_str or len(color_str) < 6:
        return None
    # Strip alpha channel (xlsxwriter handles it internally)
    if len(color_str) == 8:
        return '#' + color_str[2:]
    return '#' + color_str


def _border_style_to_int(style):
    """Convert border style string to xlsxwriter border constant."""
    if style is None:
        return 0
    mapping = {
        'thin': 1,
        'medium': 2,
        'thick': 5,
        'double': 6,
        'hair': 7,
    }
    return mapping.get(style, 0)


class ReportFormats:
    """Pre-created format objects for the entire workbook. Created once, reused by reference."""

    def __init__(self, workbook):
        self.wb = workbook
        self._format_cache = {}

    def _make_format(self, cell_def):
        """Create a xlsxwriter format from a template cell definition."""
        props = {}

        font = cell_def.get('font', {})
        if font:
            props['font_name'] = font.get('name', 'Trebuchet MS')
            props['font_size'] = font.get('size', 8.25)
            props['bold'] = font.get('bold', False)
            props['italic'] = font.get('italic', False)
            color = _convert_color(font.get('color'))
            if color:
                props['font_color'] = color

        alignment = cell_def.get('alignment', {})
        if alignment:
            h_align = alignment.get('horizontal', 'center')
            v_align = alignment.get('vertical', 'center')
            props['align'] = h_align
            props['valign'] = 'vcenter' if v_align == 'center' else v_align
            props['text_wrap'] = alignment.get('wrap_text', True)

        fill = cell_def.get('fill', {})
        if fill and fill.get('pattern') == 'solid':
            fg_color = _convert_color(fill.get('fg_color'))
            if fg_color and fg_color != '#FFFFFF':
                props['bg_color'] = fg_color

        border = cell_def.get('border', {})
        if border:
            if border.get('left'):
                props['left'] = _border_style_to_int(border['left'])
            if border.get('right'):
                props['right'] = _border_style_to_int(border['right'])
            if border.get('top'):
                props['top'] = _border_style_to_int(border['top'])
            if border.get('bottom'):
                props['bottom'] = _border_style_to_int(border['bottom'])

        num_fmt = cell_def.get('number_format', 'General')
        if num_fmt and num_fmt not in ('General',):
            props['num_format'] = num_fmt

        return self.wb.add_format(props)

    def get_format(self, cell_def):
        """Get or create a cached format from a cell definition."""
        # Build cache key from the parts that affect formatting
        key_parts = []
        font = cell_def.get('font', {})
        key_parts.append((font.get('name'), font.get('size'), font.get('bold'), font.get('italic'), font.get('color')))

        alignment = cell_def.get('alignment', {})
        key_parts.append((alignment.get('horizontal'), alignment.get('vertical'), alignment.get('wrap_text')))

        fill = cell_def.get('fill', {})
        key_parts.append((fill.get('pattern'), fill.get('fg_color')))

        border = cell_def.get('border', {})
        key_parts.append((border.get('left'), border.get('right'), border.get('top'), border.get('bottom')))

        key_parts.append(cell_def.get('number_format'))

        cache_key = tuple(key_parts)
        if cache_key not in self._format_cache:
            self._format_cache[cache_key] = self._make_format(cell_def)
        return self._format_cache[cache_key]

    def create_data_row_format(self):
        """Create format for data row cells WITH thin borders on all sides."""
        return self.wb.add_format({
            'font_name': 'Trebuchet MS',
            'font_size': 8,
            'bold': False,
            'italic': False,
            'font_color': '#000000',
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'left': 1,
            'right': 1,
            'top': 1,
            'bottom': 1,
        })

    def create_data_row_left_align_format(self):
        """Create format for data row cells with left alignment and borders."""
        return self.wb.add_format({
            'font_name': 'Trebuchet MS',
            'font_size': 8,
            'bold': False,
            'italic': False,
            'font_color': '#000000',
            'align': 'left',
            'valign': 'vcenter',
            'text_wrap': True,
            'left': 1,
            'right': 1,
            'top': 1,
            'bottom': 1,
        })

    def create_generated_by_format(self):
        """Create format for the 'Generated by' signature line."""
        return self.wb.add_format({
            'font_name': 'Trebuchet MS',
            'bold': True,
            'italic': True,
            'font_color': '#7F7F7F',
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
        })


def _write_template_header(ws, formats, template, value_overrides):
    """
    Write the template header section (rows 1-31) from the JSON template.
    This includes the company info, date range, and billing summary headers.
    """
    cells = template.get('cells', [])
    merged_cells = template.get('merged_cells', [])
    column_widths = template.get('column_widths', {})
    row_heights = template.get('row_heights', {})

    # Set column widths
    for col_str, width in column_widths.items():
        col_idx = int(col_str) - 1  # xlsxwriter uses 0-based column index
        ws.set_column(col_idx, col_idx, width)

    # Set row heights for template rows
    for row_str, height in row_heights.items():
        row_idx = int(row_str) - 1  # xlsxwriter uses 0-based row index
        ws.set_row(row_idx, height)

    # Build a lookup of which cells are the "start" of a merge range
    # so we know to write with merge_range() instead of write()
    merge_lookup = {}
    for mr in merged_cells:
        try:
            parts = mr.replace(':', '')
            # Parse e.g. "B6:Q6" -> find min row/col
            from openpyxl.utils import range_boundaries
            min_col, min_row, max_col, max_row = range_boundaries(mr)
            merge_lookup[(min_row, min_col)] = (min_row - 1, min_col - 1, max_row - 1, max_col - 1, mr)
        except Exception:
            pass

    # Filter header cells only (rows 1-31)
    header_cells = [c for c in cells if c.get('row', 0) <= 31]

    # Track which merge ranges we've already written
    written_merges = set()

    for cell_def in header_cells:
        row = cell_def.get('row')
        col = cell_def.get('col')
        value = cell_def.get('value', '')

        # Check for value override
        cell_ref = cell_def.get('cell', '')
        if cell_ref in value_overrides:
            value = value_overrides[cell_ref]

        fmt = formats.get_format(cell_def)

        # xlsxwriter uses 0-based indices
        r = row - 1
        c = col - 1

        # Check if this cell starts a merge range
        if (row, col) in merge_lookup:
            min_r, min_c, max_r, max_c, mr_str = merge_lookup[(row, col)]
            if mr_str not in written_merges:
                written_merges.add(mr_str)
                if min_r == max_r and min_c == max_c:
                    ws.write(r, c, value, fmt)
                else:
                    ws.merge_range(min_r, min_c, max_r, max_c, value, fmt)
        else:
            # Regular cell write (only if not part of a merge range as a non-start cell)
            is_inside_merge = False
            for (mr, mc), (mnr, mnc, mxr, mxc, _) in merge_lookup.items():
                if mnr <= r <= mxr and mnc <= c <= mxc and (mr, mc) != (row, col):
                    is_inside_merge = True
                    break
            if not is_inside_merge:
                ws.write(r, c, value, fmt)


def _write_billing_data(ws, formats, template, summary_bills, subscriber_name):
    """Write billing summary data to the worksheet."""
    products_to_bill = {
        12: ('consumer_snap_check', 'Consumer Snap Check'),
        13: ('consumer_basic_trace', 'Consumer Basic Trace'),
        14: ('consumer_basic_credit', 'Consumer Basic Credit'),
        15: ('consumer_detailed_credit', 'Consumer Detailed Credit'),
        16: ('xscore_consumer_detailed_credit', 'X-Score Consumer Detailed Credit'),
        17: ('commercial_basic_trace', 'Commercial Basic Trace'),
        18: ('commercial_detailed_credit', 'Commercial Detailed Credit'),
        20: ('enquiry_report', 'Enquiry Report'),
        22: ('consumer_dud_cheque', 'Consumer Dud Cheque'),
        23: ('commercial_dud_cheque', 'Commercial Dud Cheque'),
        25: ('director_basic_report', 'Director Basic Report'),
        26: ('director_detailed_report', 'Director Detailed Report'),
        19: ('iscore', 'iScore'),
        21: ('xscore_commercial_detailed_credit', 'X-Score Commercial Detailed Credit'),
        24: ('kyc_report', 'KYC report'),
        27: ('xscore_consumer_prime', 'Xscore Consumer Prime'),
    }

    # Build cell format lookup from template
    cells = template.get('cells', [])
    template_cells = {(c['row'], c['col']): c for c in cells}

    total_amount = Decimal('0.00')

    for row, (key, name) in products_to_bill.items():
        quantity = summary_bills.get(key, 0)
        rate = summary_bills.get(f'{key}_rate', Decimal('0.00'))
        amount = Decimal(quantity) * rate

        total_amount += amount

        r = row - 1  # 0-based

        if quantity == 0:
            ws.set_row(r, None, None, {'hidden': True})
            continue

        # Quantity (column I = 9, 0-based = 8)
        fmt_i = formats.get_format(template_cells.get((row, 9), {})) if (row, 9) in template_cells else None
        ws.write(r, 8, quantity, fmt_i)

        # Rate (column M = 13, 0-based = 12)
        fmt_m = formats.get_format(template_cells.get((row, 13), {})) if (row, 13) in template_cells else None
        ws.write(r, 12, f"₦{rate:,.2f}", fmt_m)

        # Amount (column P = 16, 0-based = 15)
        fmt_p = formats.get_format(template_cells.get((row, 16), {})) if (row, 16) in template_cells else None
        ws.write(r, 15, f"₦{amount:,.2f}", fmt_p)

    # Totals
    vat_amount = total_amount * Decimal('0.075')
    amount_due = total_amount + vat_amount

    # Total amount (row 28, col P)
    fmt_28 = formats.get_format(template_cells.get((28, 16), {})) if (28, 16) in template_cells else None
    ws.write(27, 15, f"₦{total_amount:,.2f}", fmt_28)

    # VAT (row 29, col P)
    fmt_29 = formats.get_format(template_cells.get((29, 16), {})) if (29, 16) in template_cells else None
    ws.write(28, 15, f"₦{vat_amount:,.2f}", fmt_29)

    # Amount due (row 30, col P)
    fmt_30 = formats.get_format(template_cells.get((30, 16), {})) if (30, 16) in template_cells else None
    ws.write(29, 15, f"₦{amount_due:,.2f}", fmt_30)


def _write_product_section_header(ws, formats, template, start_row, product_name):
    """
    Write a product section header (replicating template rows 32-35) at start_row.

    Args:
        ws: xlsxwriter worksheet
        formats: ReportFormats instance
        template: template dict
        start_row: 1-based row to start writing the header
        product_name: actual product name to replace 'Productname'
    """
    cells = template.get('cells', [])
    merged_cells = template.get('merged_cells', [])
    row_heights = template.get('row_heights', {})

    row_offset = start_row - 32  # offset from template position

    # Set row heights for headers
    for template_row in range(32, 36):
        h = row_heights.get(str(template_row))
        if h:
            ws.set_row(template_row + row_offset - 1, h)  # 0-based

    # Find header cells (rows 32-35) and corresponding merges
    header_cells = [c for c in cells if 32 <= c.get('row', 0) <= 35]

    # Build merge lookup for header rows
    from openpyxl.utils import range_boundaries
    header_merge_lookup = {}
    for mr in merged_cells:
        try:
            min_col, min_row, max_col, max_row = range_boundaries(mr)
            if 32 <= min_row <= 35:
                # Offset to target position
                new_min_row = min_row + row_offset
                new_max_row = max_row + row_offset
                header_merge_lookup[(min_row, min_col)] = (
                    new_min_row - 1, min_col - 1,
                    new_max_row - 1, max_col - 1
                )
        except Exception:
            pass

    written_merges = set()

    for cell_def in header_cells:
        row = cell_def.get('row')
        col = cell_def.get('col')
        value = cell_def.get('value', '')

        # Replace product placeholders
        if value and 'productname' in str(value).lower():
            if 'this section' in str(value).lower():
                value = f"This section includes the enquiries for the {product_name}"
            else:
                value = product_name

        target_row = row + row_offset
        r = target_row - 1  # 0-based
        c = col - 1

        fmt = formats.get_format(cell_def)

        # Check if this cell starts a merge
        if (row, col) in header_merge_lookup:
            min_r, min_c, max_r, max_c = header_merge_lookup[(row, col)]
            merge_key = (min_r, min_c, max_r, max_c)
            if merge_key not in written_merges:
                written_merges.add(merge_key)
                if min_r == max_r and min_c == max_c:
                    ws.write(r, c, value, fmt)
                else:
                    ws.merge_range(min_r, min_c, max_r, max_c, value, fmt)
        else:
            # Check not inside another merge
            is_inside_merge = False
            for (mr, mc), (mnr, mnc, mxr, mxc) in header_merge_lookup.items():
                if mnr <= r <= mxr and mnc <= c <= mxc and (mr, mc) != (row, col):
                    is_inside_merge = True
                    break
            if not is_inside_merge:
                ws.write(r, c, value, fmt)


def _setup_overflow_sheet(wb, template, sheet_number):
    """
    Create a new overflow worksheet with column widths and hidden gridlines.
    Also inserts the FCB logo at the same position as sheet 1.

    Returns:
        The new xlsxwriter worksheet
    """
    ws = wb.add_worksheet(f'Sheet{sheet_number}')
    ws.hide_gridlines(2)

    column_widths = template.get('column_widths', {})
    for col_str, width in column_widths.items():
        col_idx = int(col_str) - 1
        ws.set_column(col_idx, col_idx, width)

    # Insert FCB logo at the same position as the first sheet
    logo_path = os.path.join(settings.BASE_DIR.parent, 'media', 'Images', 'FirstCentralAPPROVEDLogo.png')
    if os.path.exists(logo_path):
        try:
            ws.insert_image('B1', logo_path, {
                'x_scale': 0.5,
                'y_scale': 0.5,
                'x_offset': 2,
                'y_offset': 2,
            })
        except Exception:
            pass

    return ws


def _write_single_data_row(ws, r, record, serial_number, data_fmt, data_left_fmt):
    """Write one data row at 0-based row index `r`."""
    ws.set_row(r, 25)

    # Format dates
    subscriber_enquiry_date_str = ""
    if record.get('SubscriberEnquiryDate'):
        enq_date = record['SubscriberEnquiryDate']
        if isinstance(enq_date, str):
            subscriber_enquiry_date_str = enq_date
        else:
            date_only = enq_date.date() if hasattr(enq_date, 'date') else enq_date
            subscriber_enquiry_date_str = date_only.strftime('%Y-%m-%d')

    details_viewed_date_str = ""
    if record.get('DetailsViewedDate'):
        view_date = record['DetailsViewedDate']
        if isinstance(view_date, str):
            details_viewed_date_str = view_date
        else:
            date_only = view_date.date() if hasattr(view_date, 'date') else view_date
            details_viewed_date_str = date_only.strftime('%Y-%m-%d')

    ws.write(r, 1, serial_number, data_fmt)
    ws.write(r, 2, "", data_fmt)
    ws.write(r, 3, "", data_fmt)

    subscriber_name = record.get('SubscriberName', '')
    ws.merge_range(r, 4, r, 5, subscriber_name, data_left_fmt)

    system_user = record.get('SystemUser', '') or ""
    ws.merge_range(r, 6, r, 8, system_user, data_fmt)

    ws.write(r, 9, subscriber_enquiry_date_str, data_fmt)
    ws.write(r, 10, record.get('ProductName', ''), data_left_fmt)
    ws.merge_range(r, 11, r, 13, details_viewed_date_str, data_fmt)

    search_output = record.get('SearchOutput', '') or ""
    ws.merge_range(r, 14, r, 16, search_output, data_left_fmt)


def _write_data_rows(
    ws, data_fmt, data_left_fmt,
    start_row, product_records, serial_number_base,
    wb=None, formats=None, template=None, product_name=None, sheet_counter=1,
):
    """
    Write product data rows with borders on ALL cells.
    Handles multi-sheet overflow: if a row would exceed MAX_ROWS_PER_SHEET,
    a new sheet is created and writing continues there.

    Args:
        ws: xlsxwriter worksheet (active)
        data_fmt: center-aligned data format with borders
        data_left_fmt: left-aligned data format with borders
        start_row: 1-based first data row
        product_records: list of record dicts
        serial_number_base: base for serial number calculation
        wb: xlsxwriter Workbook (needed for overflow)
        formats: ReportFormats instance (needed for overflow headers)
        template: template dict (needed for overflow column widths)
        product_name: current product name (used for overflow header)
        sheet_counter: current sheet number for naming

    Returns:
        (active_worksheet, final_current_row_1based, sheet_counter)
    """
    current_ws = ws

    for idx, record in enumerate(product_records):
        current_row = start_row + idx  # 1-based
        r = current_row - 1  # 0-based

        # Check if we've exceeded the sheet limit
        if current_row > MAX_ROWS_PER_SHEET and wb is not None:
            sheet_counter += 1
            current_ws = _setup_overflow_sheet(wb, template, sheet_counter)

            # Write a product header on the new sheet so the user knows context
            if formats is not None and template is not None and product_name:
                _write_product_section_header(current_ws, formats, template, 1, product_name)
                new_start = 5  # After 4 header rows
            else:
                new_start = 1

            # Recalculate: remaining records start at new_start on the new sheet
            remaining = product_records[idx:]
            result_ws, result_row, result_counter = _write_data_rows(
                current_ws, data_fmt, data_left_fmt,
                new_start, remaining, new_start - 1,
                wb=wb, formats=formats, template=template,
                product_name=product_name, sheet_counter=sheet_counter,
            )
            return result_ws, result_row, result_counter

        serial_number = current_row - serial_number_base
        _write_single_data_row(current_ws, r, record, serial_number, data_fmt, data_left_fmt)

    final_row = start_row + len(product_records)
    return current_ws, final_row, sheet_counter


def generate_full_report(
    subscriber_name,
    start_date_display,
    end_date_display,
    summary_bills=None,
    product_sections=None,
    username="",
    output_path=None,
):
    """
    Generate a complete Excel report in a single streaming pass using xlsxwriter.

    Args:
        subscriber_name: Name of the subscriber
        start_date_display: Formatted start date string
        end_date_display: Formatted end date string
        summary_bills: Dict with product counts and rates (None to skip billing)
        product_sections: Dict mapping product names to record lists (None to skip)
        username: Username for the 'Generated by' footer
        output_path: Full path to write the .xlsx file

    Returns:
        The output_path where the file was written
    """
    import time

    if output_path is None:
        fd, output_path = tempfile.mkstemp(suffix='.xlsx')
        os.close(fd)

    template = load_template_structure()

    # Create workbook with optimizations
    wb = xlsxwriter.Workbook(output_path, {
        'constant_memory': False,  # Need random access for merged cells
        'strings_to_urls': False,  # Don't auto-convert strings to URLs
        'default_date_format': 'yyyy-mm-dd',
    })

    ws = wb.add_worksheet('Sheet1')
    ws.hide_gridlines(2)  # Hide screen and printed gridlines for a clean dashboard look

    # Pre-create ALL format objects (done once, reused by reference)
    formats = ReportFormats(wb)
    data_fmt = formats.create_data_row_format()
    data_left_fmt = formats.create_data_row_left_align_format()
    generated_by_fmt = formats.create_generated_by_format()

    # Prepare dynamic values
    value_overrides = prepare_dynamic_values(subscriber_name, start_date_display, end_date_display)

    # Phase 1: Write template header (rows 1-31)
    t0 = time.time()
    _write_template_header(ws, formats, template, value_overrides)
    print(f"   [xlsxwriter] Header written in {time.time() - t0:.2f}s")

    # Phase 2: Write billing data (rows 11-30)
    if summary_bills:
        t1 = time.time()
        _write_billing_data(ws, formats, template, summary_bills, subscriber_name)
        print(f"   [xlsxwriter] Billing data written in {time.time() - t1:.2f}s")

    # Phase 3: Write product sections
    current_row = 36  # Start after template header
    sheet_counter = 1  # Track sheet numbers for overflow
    if product_sections:
        sorted_product_names = sorted(product_sections.keys())

        for product_idx, product_name in enumerate(sorted_product_names):
            product_records = product_sections[product_name]
            t2 = time.time()

            # Check if we need to overflow to a new sheet before this product
            if current_row + 4 > MAX_ROWS_PER_SHEET:
                sheet_counter += 1
                ws = _setup_overflow_sheet(wb, template, sheet_counter)
                current_row = 1

            if product_idx == 0 and sheet_counter == 1:
                # First product uses the existing template header (rows 32-35)
                _write_product_section_header(ws, formats, template, 32, product_name)
                data_start_row = 36
                serial_number_base = 35
            else:
                # Add spacing between products
                current_row += 4

                # Clone product section header at current_row
                _write_product_section_header(ws, formats, template, current_row, product_name)
                data_start_row = current_row + 4  # After 4 header rows (32-35)
                serial_number_base = data_start_row - 1

            # Write data rows with borders on ALL cells (handles overflow internally)
            ws, current_row, sheet_counter = _write_data_rows(
                ws, data_fmt, data_left_fmt,
                data_start_row, product_records, serial_number_base,
                wb=wb, formats=formats, template=template,
                product_name=product_name, sheet_counter=sheet_counter,
            )

            elapsed = time.time() - t2
            print(f"   [xlsxwriter] Product '{product_name}' ({len(product_records)} records) in {elapsed:.2f}s")

    # Phase 4: Add logo
    logo_path = os.path.join(settings.BASE_DIR.parent, 'media', 'Images', 'FirstCentralAPPROVEDLogo.png')
    if os.path.exists(logo_path):
        try:
            ws.insert_image('B1', logo_path, {
                'x_scale': 0.5,
                'y_scale': 0.5,
                'x_offset': 2,
                'y_offset': 2,
            })
        except Exception as e:
            logger.warning(f"Failed to add logo: {e}")

    # Phase 5: Generated by footer
    if username:
        sig_row = current_row + 1  # Two rows below last data (0-based = current_row)
        ws.merge_range(sig_row, 14, sig_row, 16, f"Report Generated by: {username}", generated_by_fmt)
        ws.set_row(sig_row, 26)

    # Close workbook (this triggers the actual write to disk)
    t_save = time.time()
    wb.close()
    print(f"   [xlsxwriter] File saved in {time.time() - t_save:.2f}s")

    return output_path


# ============================================================================
# BACKWARD COMPATIBILITY - Keep these for any code that still references them
# ============================================================================

def create_workbook_from_json(value_overrides=None):
    """Legacy function - generates a report using xlsxwriter now."""
    raise NotImplementedError(
        "create_workbook_from_json() is deprecated. Use generate_full_report() instead."
    )

def write_billing_data(ws, summary_bills, subscriber_name, logger):
    """Legacy function - billing is now handled inside generate_full_report()."""
    raise NotImplementedError(
        "write_billing_data() is deprecated. Use generate_full_report() instead."
    )

def optimized_write_product_data(ws, start_row, product_records, serial_number_base, chunk_size=1000):
    """Legacy function - data writing is now handled inside generate_full_report()."""
    raise NotImplementedError(
        "optimized_write_product_data() is deprecated. Use generate_full_report() instead."
    )
