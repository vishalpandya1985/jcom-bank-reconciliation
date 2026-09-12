from reconcile_engine import load_gl, load_hdfc, reconcile
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

FONT = 'Arial'
HEADER_FILL = PatternFill('solid', fgColor='1F4E78')
HEADER_FONT = Font(name=FONT, bold=True, color='FFFFFF', size=10)
TITLE_FONT = Font(name=FONT, bold=True, size=14, color='1F4E78')
SUB_FONT = Font(name=FONT, size=10, italic=True, color='555555')
NORMAL_FONT = Font(name=FONT, size=10)
BOLD_FONT = Font(name=FONT, size=10, bold=True)

MATCH_FILL = PatternFill('solid', fgColor='C6EFCE')
PENDING_FILL = PatternFill('solid', fgColor='FFEB9C')
NA_FILL = PatternFill('solid', fgColor='D9D9D9')
MULTI_FILL = PatternFill('solid', fgColor='FFC7CE')
EXCLUDED_FILL = PatternFill('solid', fgColor='D9D2E9')

THIN = Side(style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

GL_DESCRIPTIONS = {
    '3493': 'Outward RTGS/NEFT',
    '3496': 'Inward RTGS/NEFT',
    '345051': 'IMPS/UPI/POS/ACH/NACH/Returns',
}


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = BORDER


def autosize(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def status_fill(status):
    if status.startswith('Matched'):
        return MATCH_FILL
    if status.startswith('Excluded'):
        return EXCLUDED_FILL
    if status == 'Not Applicable - Internal GL Sweep':
        return NA_FILL
    if status == 'Pending - Multiple Possible Matches':
        return MULTI_FILL
    return PENDING_FILL


def _serialize_gl_rows(rows, hdfc_rows):
    out = []
    for row in rows:
        matched_ref = ''
        matched_date = ''
        if row['matched_hdfc_row'] is not None:
            h = hdfc_rows[row['matched_hdfc_row']]
            matched_ref = h['Reference No']
            matched_date = h['Transaction Date']
        out.append({
            'post_date': row['Post Date'],
            'value_date': row['Value Date'],
            'narration': row['Narration'],
            'dr_cr': row['Dr/Cr'],
            'amount': row['Amount'],
            'status': row['status'],
            'matched_ref': matched_ref,
            'matched_date': matched_date,
            'note': row['note'],
        })
    return out


def _serialize_hdfc_unmatched(hdfc_rows):
    out = []
    for h in hdfc_rows:
        if h['used']:
            continue
        out.append({
            'date': h['Transaction Date'],
            'description': h['Description'],
            'amount': h['Amount'],
            'dr_cr': h['Dr/Cr'],
            'ref': h['Reference No'],
            'branch': h['Branch'],
        })
    return out


def build_workbook_from_detail(detail):
    """Build the multi-sheet Excel workbook purely from the already-serialized
    detail dict (per_gl entries + hdfc_unmatched). Used both for the initial
    reconciliation run and for regenerating the report after a 're-check
    pending against a later statement' update, since both operate on the
    same JSON-shaped data once the raw in-memory GL/HDFC objects are gone.
    Returns (workbook, summary)."""
    wb = Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet('Summary')
    ws['A1'] = 'RTGS / NEFT / IMPS / UPI / ACH / POS Reconciliation Report'
    ws['A1'].font = TITLE_FONT
    ws['A2'] = 'THE JUNAGADH COMM.CO.OP.BANK LTD vs Internal GL Suspense Heads'
    ws['A2'].font = SUB_FONT
    ws['A3'] = 'Matching basis: exact match on amount + reference/transaction code/name. Ambiguous cases are left Pending. Nothing is deleted.'
    ws['A3'].font = SUB_FONT
    ws.merge_cells('A1:F1')
    ws.merge_cells('A2:F2')
    ws.merge_cells('A3:F3')

    headers = ['GL Head', 'Description', 'Total Entries', 'Matched', 'Pending', 'Internal Sweep (N/A)']
    r = 5
    for c, h in enumerate(headers, start=1):
        ws.cell(row=r, column=c, value=h)
    style_header(ws, r, len(headers))

    r += 1
    summary = {'per_gl': {}, 'total': 0, 'matched': 0, 'pending': 0, 'na': 0, 'excluded': 0, 'hdfc_unmatched': 0}
    for gl_code in ['3493', '3496', '345051']:
        rows = detail['per_gl'][gl_code]
        desc = GL_DESCRIPTIONS[gl_code]
        matched = sum(1 for x in rows if x['status'].startswith('Matched'))
        na = sum(1 for x in rows if x['status'] == 'Not Applicable - Internal GL Sweep')
        excluded = sum(1 for x in rows if x['status'].startswith('Excluded'))
        pending = len(rows) - matched - na - excluded
        ws.cell(row=r, column=1, value=gl_code).font = BOLD_FONT
        ws.cell(row=r, column=2, value=desc).font = NORMAL_FONT
        ws.cell(row=r, column=3, value=len(rows)).font = NORMAL_FONT
        ws.cell(row=r, column=4, value=matched).font = NORMAL_FONT
        ws.cell(row=r, column=5, value=pending).font = NORMAL_FONT
        ws.cell(row=r, column=6, value=na).font = NORMAL_FONT
        for c in range(1, 7):
            ws.cell(row=r, column=c).border = BORDER
        summary['per_gl'][gl_code] = {
            'description': desc, 'total': len(rows), 'matched': matched, 'pending': pending,
            'na': na, 'excluded': excluded
        }
        summary['total'] += len(rows)
        summary['matched'] += matched
        summary['pending'] += pending
        summary['na'] += na
        summary['excluded'] += excluded
        r += 1

    ws.cell(row=r, column=1, value='TOTAL').font = BOLD_FONT
    ws.cell(row=r, column=3, value=summary['total']).font = BOLD_FONT
    ws.cell(row=r, column=4, value=summary['matched']).font = BOLD_FONT
    ws.cell(row=r, column=5, value=summary['pending']).font = BOLD_FONT
    ws.cell(row=r, column=6, value=summary['na']).font = BOLD_FONT
    for c in range(1, 7):
        ws.cell(row=r, column=c).border = BORDER
    r += 2

    unmatched_hdfc = len(detail.get('hdfc_unmatched', []))
    summary['hdfc_unmatched'] = unmatched_hdfc
    ws.cell(row=r, column=1, value='HDFC statement lines with no GL match:').font = BOLD_FONT
    ws.cell(row=r, column=4, value=unmatched_hdfc).font = BOLD_FONT
    r += 1
    ws.cell(row=r, column=1, value='(See "HDFC Unmatched" tab)').font = SUB_FONT

    autosize(ws, [14, 40, 14, 10, 10, 18])

    gl_headers = ['Post Date', 'Value Date', 'Narration', 'Dr/Cr', 'Amount', 'Status', 'Matched HDFC Ref No', 'Matched HDFC Date', 'Note']
    for gl_code in ['3493', '3496', '345051']:
        rows = detail['per_gl'][gl_code]
        desc = GL_DESCRIPTIONS[gl_code]
        ws = wb.create_sheet(f'GL {gl_code}')
        ws['A1'] = f'GL {gl_code} - {desc}'
        ws['A1'].font = TITLE_FONT
        ws.merge_cells('A1:I1')
        r0 = 3
        for c, h in enumerate(gl_headers, start=1):
            ws.cell(row=r0, column=c, value=h)
        style_header(ws, r0, len(gl_headers))

        r = r0 + 1
        for row in rows:
            values = [row['post_date'], row['value_date'], row['narration'], row['dr_cr'],
                      row['amount'], row['status'], row['matched_ref'], row['matched_date'], row['note']]
            for c, v in enumerate(values, start=1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.font = NORMAL_FONT
                cell.border = BORDER
                cell.fill = status_fill(row['status'])
                if c == 5:
                    cell.number_format = '#,##0.00'
            r += 1
        autosize(ws, [11, 11, 55, 7, 14, 30, 20, 12, 42])
        ws.freeze_panes = 'A4'

    ws = wb.create_sheet('HDFC Unmatched')
    ws['A1'] = 'HDFC Statement Lines With No GL Match'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:H1')
    hh = ['Transaction Date', 'Description', 'Amount', 'Dr/Cr', 'Reference No', 'Value Date', 'Branch', 'Running Balance']
    r0 = 3
    for c, h in enumerate(hh, start=1):
        ws.cell(row=r0, column=c, value=h)
    style_header(ws, r0, len(hh))
    r = r0 + 1
    for h in detail.get('hdfc_unmatched', []):
        values = [h['date'], h['description'], h['amount'], h['dr_cr'], h['ref'], '', h['branch'], '']
        for c, v in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = NORMAL_FONT
            cell.border = BORDER
            cell.fill = PENDING_FILL
            if c == 3:
                cell.number_format = '#,##0.00'
        r += 1
    autosize(ws, [18, 55, 14, 7, 20, 12, 22, 16])
    ws.freeze_panes = 'A4'

    wb.move_sheet('Summary', offset=-len(wb.sheetnames))
    return wb, summary


def run_reconciliation(file_paths, output_xlsx_path):
    """
    file_paths: dict with keys '3493', '3496', '345051', 'hdfc' -> filesystem paths
    output_xlsx_path: where to save the generated .xlsx report
    Returns: (summary dict, detail dict) - detail has per-GL row-level entries
             plus the unmatched HDFC lines, for on-screen display.
    """
    hdfc_rows = load_hdfc(file_paths['hdfc'])

    gl_data = {}
    for gl_code in ['3493', '3496', '345051']:
        gl_data[gl_code] = load_gl(file_paths[gl_code])

    for gl_code in ['3496', '3493', '345051']:
        reconcile(gl_data[gl_code], hdfc_rows)

    detail = {
        'per_gl': {gl_code: _serialize_gl_rows(gl_data[gl_code], hdfc_rows) for gl_code in ['3493', '3496', '345051']},
        'hdfc_unmatched': _serialize_hdfc_unmatched(hdfc_rows),
    }

    wb, summary = build_workbook_from_detail(detail)
    wb.save(output_xlsx_path)

    return summary, detail
